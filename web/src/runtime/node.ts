// Just enough of Node.js for the debug build of a RAS program (`ras build --target js`) to run
// in the browser: CommonJS modules read from a `MemFs`, `process`, and the built-in modules the
// run time asks for (`fs`, `path`, `os`, `util`, `module`, ...). The `ras_book` addon is the
// WebAssembly `Book` (crates/ras-wasm), given as the module `ras_book`.

import { Buffer } from "buffer";
import * as pathModule from "path-browserify";
import { SourceMapConsumer } from "source-map-js";
import { MemFs } from "./memfs";

/** V8's frames (`Error.prepareStackTrace`). */
interface CallSite {
    getFileName(): string | null;
    getScriptNameOrSourceURL?(): string | null;
    getLineNumber(): number | null;
    getColumnNumber(): number | null;
    getFunctionName(): string | null;
    getMethodName(): string | null;
    isConstructor(): boolean;
}

export type Output = (stream: "stdout" | "stderr", text: string) => void;

/** What the program's `process.exit()` throws to stop it. */
export class ProcessExit extends Error {
    constructor(readonly code: number) {
        super(`process.exit(${code})`);
    }
}

function inspect(v: unknown, depth = 2): string {
    if (typeof v === "string") return v;
    if (v instanceof Error) return v.stack || `${v.name}: ${v.message}`;
    if (v === null || typeof v !== "object") return String(v);
    try {
        const seen = new WeakSet();
        return JSON.stringify(
            v,
            (_k, x) => {
                if (typeof x === "object" && x !== null) {
                    if (seen.has(x)) return "[Circular]";
                    seen.add(x);
                }
                return typeof x === "bigint" ? `${x}n` : x;
            },
            depth > 0 ? 2 : undefined,
        );
    } catch {
        return String(v);
    }
}

function format(args: unknown[]): string {
    if (typeof args[0] === "string" && /%[sdifjoO%]/.test(args[0])) {
        let i = 1;
        const head = args[0].replace(/%([sdifjoO%])/g, (m, k) => {
            if (k === "%") return "%";
            if (i >= args.length) return m;
            const a = args[i++];
            if (k === "d" || k === "i") return String(parseInt(String(a), 10));
            if (k === "f") return String(parseFloat(String(a)));
            if (k === "s") return typeof a === "string" ? a : inspect(a, 0);
            return inspect(a);
        });
        return [head, ...args.slice(i).map((a) => inspect(a))].join(" ");
    }
    return args.map((a) => inspect(a)).join(" ");
}

export interface NodeOptions {
    fs: MemFs;
    out: Output;
    env: Record<string, string>;
    argv: string[];
    /** Modules that are not files: `ras_book`. */
    extra: Record<string, unknown>;
    /** Called with each module of `fs` once it has loaded. */
    loaded?: (id: string, exports: unknown) => void;
}

/**
 * A module of the debug build, ready for the browser's DevTools to show its `.ras` files
 * (spec §14.10.9): the `sourceMappingURL` of the build names the `.map` next to the module, which
 * is only in `fs`, and the map names the `.ras` files without their text. The map is inlined as a
 * `data:` URL with the sources' text (`sourcesContent`; the standard library's is in the map
 * already), and the sources become `ras://` URLs
 * of their paths in `fs` (`ras://project/src/main.ras`), absolute whatever the page's URL is.
 * `body` is the module without its `sourceMappingURL`; `sourceMap` the comment to end the
 * evaluated code with (empty when the module has no map).
 */
export function inlineSourceMap(text: string, id: string, fs: MemFs): { body: string; sourceMap: string } {
    const m = /\/\/# sourceMappingURL=([^\s]+)\s*$/.exec(text);
    if (!m || m[1].startsWith("data:")) return { body: text, sourceMap: "" };
    const body = text.slice(0, m.index);
    const dir = pathModule.dirname(id);
    let map: { sources?: string[]; sourcesContent?: (string | null)[]; sourceRoot?: string };
    try {
        map = JSON.parse(fs.readFileSync(pathModule.resolve(dir, m[1]), "utf8") as string);
    } catch {
        return { body, sourceMap: "" };
    }
    const files = (map.sources || []).map((s) => pathModule.resolve(dir, map.sourceRoot || "", s));
    // `ras://project/src/main.ras`; the standard library (`<stdlib>/…`) is `ras://stdlib/…`.
    map.sources = files.map((f) => `ras://${f.replace(/^.*\/<stdlib>\//, "stdlib/").replace(/^\/+/, "")}`);
    // The standard library is not in `fs`: its text comes in the map (`sourcesContent`).
    const given = map.sourcesContent || [];
    map.sourcesContent = files.map((f, i) => {
        try {
            return fs.readFileSync(f, "utf8") as string;
        } catch {
            return given[i] ?? null;
        }
    });
    delete map.sourceRoot;
    const data = Buffer.from(JSON.stringify(map), "utf8").toString("base64");
    return { body, sourceMap: `\n//# sourceMappingURL=data:application/json;charset=utf-8;base64,${data}` };
}

/** A Node-like world: `run(file)` loads a module of `fs` as Node runs a script. */
export class NodeWorld {
    readonly fs: MemFs;
    readonly process: Record<string, unknown> & { exitCode?: number; env: Record<string, string> };
    private cache = new Map<string, { exports: unknown }>();
    private handlers = new Map<string, ((...a: unknown[]) => void)[]>();
    private builtins: Record<string, unknown>;
    private maps = new Map<string, SourceMapConsumer | null>();

    constructor(private opts: NodeOptions) {
        this.fs = opts.fs;
        const world = this;
        const out = opts.out;
        const writer = (stream: "stdout" | "stderr") => ({
            write(chunk: unknown) {
                out(stream, typeof chunk === "string" ? chunk : Buffer.from(chunk as Uint8Array).toString("utf8"));
                return true;
            },
            isTTY: false,
            columns: 120,
        });
        this.process = {
            env: opts.env,
            argv: ["node", ...opts.argv],
            platform: "linux",
            arch: "wasm32",
            pid: 1,
            version: "v22.0.0",
            versions: { node: "22.0.0" },
            execPath: "/usr/bin/node",
            exitCode: undefined,
            stdout: writer("stdout"),
            stderr: writer("stderr"),
            cwd: () => opts.fs.cwd,
            chdir: () => {},
            on(event: string, f: (...a: unknown[]) => void) {
                const list = world.handlers.get(event) || [];
                list.push(f);
                world.handlers.set(event, list);
                return this;
            },
            once(event: string, f: (...a: unknown[]) => void) {
                return (this as unknown as { on: (e: string, f: unknown) => unknown }).on(event, f);
            },
            off(event: string, f: unknown) {
                const list = world.handlers.get(event) || [];
                world.handlers.set(
                    event,
                    list.filter((g) => g !== f),
                );
                return this;
            },
            removeListener(event: string, f: unknown) {
                return (this as unknown as { off: (e: string, f: unknown) => unknown }).off(event, f);
            },
            emitWarning: (w: unknown) => out("stderr", `${inspect(w)}\n`),
            exit: (code?: number) => {
                throw new ProcessExit(code ?? (world.process.exitCode as number) ?? 0);
            },
            hrtime: Object.assign(
                (prev?: [number, number]) => {
                    const t = performance.now();
                    const s = Math.floor(t / 1000);
                    const ns = Math.floor((t % 1000) * 1e6);
                    return prev ? [s - prev[0], ns - prev[1]] : [s, ns];
                },
                { bigint: () => BigInt(Math.floor(performance.now() * 1e6)) },
            ),
            memoryUsage: () => ({ rss: 0, heapTotal: 0, heapUsed: 0, external: 0 }),
            nextTick: (f: (...a: unknown[]) => void, ...a: unknown[]) => queueMicrotask(() => f(...a)),
        };
        const unavailable = (name: string) =>
            new Proxy(
                {},
                {
                    get(_t, prop) {
                        if (prop === "__esModule" || typeof prop === "symbol") return undefined;
                        return () => {
                            throw new Error(`${name}.${String(prop)} はブラウザでは使えません`);
                        };
                    },
                },
            );
        this.builtins = {
            fs: opts.fs,
            path: pathModule,
            os: {
                userInfo: () => ({ username: "user", homedir: "/home/user" }),
                tmpdir: () => "/tmp",
                homedir: () => "/home/user",
                hostname: () => "browser",
                platform: () => "linux",
                type: () => "Linux",
                release: () => "",
                EOL: "\n",
                cpus: () => [],
            },
            util: { inspect: (v: unknown) => inspect(v), format: (...a: unknown[]) => format(a), TextDecoder, TextEncoder },
            module: { findSourceMap: (file: string) => this.findSourceMap(file) },
            buffer: { Buffer },
            child_process: unavailable("child_process"),
            worker_threads: unavailable("worker_threads"),
            net: unavailable("net"),
            ...opts.extra,
        };
    }

    /** Node's `findSourceMap` over the `.map` files next to the modules. */
    private findSourceMap(file: string) {
        const abs = this.fs.abs(file);
        if (!this.maps.has(abs)) {
            let consumer: SourceMapConsumer | null = null;
            try {
                const text = this.fs.readFileSync(`${abs}.map`, "utf8") as string;
                consumer = new SourceMapConsumer(JSON.parse(text));
            } catch {
                consumer = null;
            }
            this.maps.set(abs, consumer);
        }
        const map = this.maps.get(abs);
        if (!map) return undefined;
        const dir = pathModule.dirname(abs);
        return {
            findEntry(line: number, column: number) {
                const p = map.originalPositionFor({ line: line + 1, column, bias: SourceMapConsumer.GREATEST_LOWER_BOUND });
                if (p.source === null || p.line === null) {
                    const q = map.originalPositionFor({ line: line + 1, column, bias: SourceMapConsumer.LEAST_UPPER_BOUND });
                    if (q.source === null || q.line === null) return {};
                    return { originalSource: pathModule.resolve(dir, q.source), originalLine: q.line - 1, originalColumn: q.column ?? 0 };
                }
                return { originalSource: pathModule.resolve(dir, p.source), originalLine: p.line - 1, originalColumn: p.column ?? 0 };
            },
        };
    }

    private resolve(request: string, fromDir: string): string {
        if (request.startsWith("node:")) request = request.slice(5);
        if (request in this.builtins) return request;
        const base = request.startsWith("/") ? request : pathModule.resolve(fromDir, request);
        for (const p of [base, `${base}.js`, `${base}.json`, pathModule.join(base, "index.js")]) {
            try {
                if (this.fs.statSync(p)?.isFile()) return p;
            } catch {
                // the next
            }
        }
        const e = new Error(`Cannot find module '${request}'`) as Error & { code: string };
        e.code = "MODULE_NOT_FOUND";
        throw e;
    }

    require(request: string, fromDir: string): unknown {
        const id = this.resolve(request, fromDir);
        if (id in this.builtins) return this.builtins[id];
        const cached = this.cache.get(id);
        if (cached) return cached.exports;
        const module = { exports: {} as unknown, id, filename: id, loaded: false };
        this.cache.set(id, module);
        try {
            const text = this.fs.readFileSync(id, "utf8") as string;
            if (id.endsWith(".json")) {
                module.exports = JSON.parse(text);
            } else {
                const dir = pathModule.dirname(id);
                const req = Object.assign((r: string) => this.require(r, dir), {
                    resolve: (r: string) => this.resolve(r, dir),
                    cache: {},
                });
                // On one line with the module's first line, so that its lines and columns are
                // the file's, as in Node (the stack traces are mapped back to the .ras files).
                const { body, sourceMap } = inlineSourceMap(text.replace(/^#!.*/, ""), id, this.fs);
                const wrapper = (0, eval)(`(function (exports, require, module, __filename, __dirname, process, Buffer) {${body}\n})\n//# sourceURL=${id}${sourceMap}`);
                wrapper.call(module.exports, module.exports, req, module, id, dir, this.process, Buffer);
            }
            module.loaded = true;
            if (this.opts.loaded) this.opts.loaded(id, module.exports);
        } catch (e) {
            this.cache.delete(id);
            throw e;
        }
        return module.exports;
    }

    /**
     * The stack traces of Node with `--enable-source-maps` (V8's `Error.prepareStackTrace`, so
     * Chromium-based browsers): the frames of the program name the `.ras` files, where the run
     * time and the test runner read the position of an error.
     */
    private mappedStack = (err: Error, frames: CallSite[]): string => {
        let head: string;
        try {
            head = String(err);
        } catch {
            head = "Error";
        }
        const lines = frames.map((f) => {
            // (The modules are evaluated: their file is the `sourceURL`.)
            const file = (f.getScriptNameOrSourceURL && f.getScriptNameOrSourceURL()) || f.getFileName();
            let line = f.getLineNumber();
            let col = f.getColumnNumber();
            let where = file || "<anonymous>";
            if (file && line !== null && col !== null && file.startsWith("/")) {
                const map = this.findSourceMap(file);
                const e = map && map.findEntry(line - 1, col - 1);
                if (e && e.originalSource) {
                    where = e.originalSource;
                    line = e.originalLine + 1;
                    col = e.originalColumn + 1;
                }
            }
            const pos = line === null ? where : `${where}:${line}:${col}`;
            const fn = f.getFunctionName() || f.getMethodName();
            const name = f.isConstructor() ? `new ${fn || "<anonymous>"}` : fn;
            return name ? `    at ${name} (${pos})` : `    at ${pos}`;
        });
        return [head, ...lines].join("\n");
    };

    emit(event: string, ...args: unknown[]): boolean {
        const list = this.handlers.get(event) || [];
        for (const f of list) f(...args);
        return list.length > 0;
    }

    /** Calls `f` with the globals of this world (`process`, `Buffer`, the mapped stacks). */
    within<T>(f: () => T): T {
        const g = globalThis as Record<string, unknown>;
        const E = Error as unknown as { prepareStackTrace?: unknown };
        const saved = { process: g.process, Buffer: g.Buffer, prepare: E.prepareStackTrace };
        g.process = this.process;
        g.Buffer = Buffer;
        E.prepareStackTrace = this.mappedStack;
        try {
            return f();
        } finally {
            g.process = saved.process;
            g.Buffer = saved.Buffer;
            E.prepareStackTrace = saved.prepare;
        }
    }

    /**
     * Calls `f` as the program's code runs: `process.exit()` and an error that escapes it (to
     * the `uncaughtException` handlers, or printed) give the exit code; undefined: it returned.
     */
    guard(f: () => void): number | undefined {
        return this.within(() => {
            try {
                f();
                return undefined;
            } catch (e) {
                if (e instanceof ProcessExit) return e.code;
                if (!this.emit("uncaughtException", e)) {
                    this.opts.out("stderr", `${inspect(e)}\n`);
                    return 1;
                }
                return undefined;
            }
        });
    }

    /** Loads a script as `node <file>` does, without ending (it may serve requests after). */
    start(file: string): number | undefined {
        return this.guard(() => this.require(file, this.fs.cwd));
    }

    /** The program ends with `code` (undefined: `process.exitCode`): the `exit` handlers run. Returns the exit code. */
    finish(code?: number): number {
        return this.within(() => {
            if (code === undefined) code = typeof this.process.exitCode === "number" ? this.process.exitCode : 0;
            this.process.exitCode = code;
            try {
                this.emit("exit", code);
            } catch (e) {
                if (!(e instanceof ProcessExit)) this.opts.out("stderr", `${inspect(e)}\n`);
            }
            return typeof this.process.exitCode === "number" ? this.process.exitCode : code;
        });
    }

    /**
     * Runs a script as `node <file>` does: an error that escapes it goes to the
     * `uncaughtException` handlers (or is printed), and then the `exit` handlers run.
     * Returns the exit code.
     */
    run(file: string): number {
        return this.finish(this.start(file));
    }
}

export { format as formatConsole };
