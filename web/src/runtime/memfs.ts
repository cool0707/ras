// An in-memory file system with the synchronous part of Node's `fs` that the RAS run time uses
// (ras-runtime.js, ras-excel.js, ras-scripting.js, the launchers). Paths are POSIX and absolute
// (`/project/src/main.ras`); relative paths are taken from `/project`.

import { Buffer } from "buffer";
import * as path from "path-browserify";

type Node = { kind: "file"; data: Uint8Array; mtime: Date } | { kind: "dir"; mtime: Date };

export class FsError extends Error {
    code: string;
    errno: number;
    path: string;
    constructor(code: string, syscall: string, p: string) {
        const text: Record<string, string> = {
            ENOENT: "no such file or directory",
            EEXIST: "file already exists",
            ENOTDIR: "not a directory",
            EISDIR: "illegal operation on a directory",
            ENOTEMPTY: "directory not empty",
            EBADF: "bad file descriptor",
        };
        super(`${code}: ${text[code] || code}, ${syscall} '${p}'`);
        this.code = code;
        this.errno = { ENOENT: -2, EEXIST: -17, ENOTDIR: -20, EISDIR: -21, ENOTEMPTY: -39, EBADF: -9 }[code] ?? -1;
        this.path = p;
    }
}

class Stats {
    constructor(private node: Node, readonly size: number) {}
    get mtime() {
        return this.node.mtime;
    }
    get mtimeMs() {
        return this.node.mtime.getTime();
    }
    get birthtime() {
        return this.node.mtime;
    }
    get ctime() {
        return this.node.mtime;
    }
    get atime() {
        return this.node.mtime;
    }
    get mode() {
        return this.node.kind === "dir" ? 0o40755 : 0o100644;
    }
    isFile() {
        return this.node.kind === "file";
    }
    isDirectory() {
        return this.node.kind === "dir";
    }
    isSymbolicLink() {
        return false;
    }
}

class Dirent {
    constructor(readonly name: string, private dir: boolean, readonly parentPath: string) {}
    isFile() {
        return !this.dir;
    }
    isDirectory() {
        return this.dir;
    }
    isSymbolicLink() {
        return false;
    }
}

function toBytes(data: unknown, encoding?: unknown): Uint8Array {
    const enc = typeof encoding === "string" ? encoding : encoding && typeof encoding === "object" ? (encoding as { encoding?: string }).encoding : undefined;
    if (typeof data === "string") return Buffer.from(data, (enc || "utf8") as "utf8");
    if (data instanceof Uint8Array) return new Uint8Array(data);
    if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength));
    if (data instanceof ArrayBuffer) return new Uint8Array(data.slice(0));
    return Buffer.from(String(data));
}

export class MemFs {
    private nodes = new Map<string, Node>();
    private fds = new Map<number, { file: string; append: boolean }>();
    private nextFd = 10;
    private tmpCount = 0;

    constructor(readonly cwd = "/project") {
        this.nodes.set("/", { kind: "dir", mtime: new Date() });
    }

    abs(p: unknown): string {
        let s = String(p instanceof URL ? p.pathname : p).replace(/\\/g, "/");
        if (s.startsWith("file://")) s = s.slice(7);
        return path.resolve(this.cwd, s);
    }

    private parentOk(p: string, syscall: string) {
        const dir = path.dirname(p);
        const parent = this.nodes.get(dir);
        if (!parent) throw new FsError("ENOENT", syscall, p);
        if (parent.kind !== "dir") throw new FsError("ENOTDIR", syscall, p);
    }

    /** All the files under a directory (for the playground): paths relative to it. */
    filesUnder(dir: string): Map<string, Uint8Array> {
        const base = this.abs(dir);
        const out = new Map<string, Uint8Array>();
        for (const [p, n] of this.nodes) {
            if (n.kind === "file" && p.startsWith(`${base}/`)) out.set(p.slice(base.length + 1), n.data);
        }
        return out;
    }

    /** Writes a file, making its directories (for the playground). */
    put(p: string, data: string | Uint8Array) {
        const a = this.abs(p);
        this.mkdirSync(path.dirname(a), { recursive: true });
        this.writeFileSync(a, data);
    }

    existsSync = (p: unknown) => {
        try {
            return this.nodes.has(this.abs(p));
        } catch {
            return false;
        }
    };

    statSync = (p: unknown, opts?: { throwIfNoEntry?: boolean }) => {
        const a = this.abs(p);
        const n = this.nodes.get(a);
        if (!n) {
            if (opts && opts.throwIfNoEntry === false) return undefined;
            throw new FsError("ENOENT", "stat", a);
        }
        return new Stats(n, n.kind === "file" ? n.data.length : 4096);
    };
    lstatSync = (p: unknown, opts?: { throwIfNoEntry?: boolean }) => this.statSync(p, opts);

    readFileSync = (p: unknown, opts?: unknown): string | Buffer => {
        if (typeof p === "number") p = this.fdFile(p, "read");
        const a = this.abs(p);
        const n = this.nodes.get(a);
        if (!n) throw new FsError("ENOENT", "open", a);
        if (n.kind === "dir") throw new FsError("EISDIR", "read", a);
        const enc = typeof opts === "string" ? opts : opts && typeof opts === "object" ? (opts as { encoding?: string }).encoding : undefined;
        const buf = Buffer.from(n.data);
        return enc ? buf.toString(enc as "utf8") : buf;
    };

    writeFileSync = (p: unknown, data: unknown, opts?: unknown) => {
        if (typeof p === "number") p = this.fdFile(p, "write");
        const a = this.abs(p);
        const n = this.nodes.get(a);
        if (n && n.kind === "dir") throw new FsError("EISDIR", "open", a);
        this.parentOk(a, "open");
        const flag = opts && typeof opts === "object" ? (opts as { flag?: string }).flag : undefined;
        const bytes = toBytes(data, opts);
        if (flag && flag.startsWith("a") && n && n.kind === "file") {
            const joined = new Uint8Array(n.data.length + bytes.length);
            joined.set(n.data);
            joined.set(bytes, n.data.length);
            this.nodes.set(a, { kind: "file", data: joined, mtime: new Date() });
        } else {
            this.nodes.set(a, { kind: "file", data: bytes, mtime: new Date() });
        }
    };

    appendFileSync = (p: unknown, data: unknown, opts?: unknown) => {
        const enc = typeof opts === "string" ? opts : undefined;
        this.writeFileSync(p, data, { encoding: enc, flag: "a" });
    };

    mkdirSync = (p: unknown, opts?: { recursive?: boolean } | number) => {
        const a = this.abs(p);
        const recursive = typeof opts === "object" && opts !== null && opts.recursive;
        const n = this.nodes.get(a);
        if (n) {
            if (recursive && n.kind === "dir") return undefined;
            throw new FsError("EEXIST", "mkdir", a);
        }
        if (recursive) {
            const parts = a.split("/").filter(Boolean);
            let cur = "";
            let first: string | undefined;
            for (const part of parts) {
                cur += `/${part}`;
                const at = this.nodes.get(cur);
                if (!at) {
                    this.nodes.set(cur, { kind: "dir", mtime: new Date() });
                    if (first === undefined) first = cur;
                } else if (at.kind !== "dir") {
                    throw new FsError("ENOTDIR", "mkdir", a);
                }
            }
            return first;
        }
        this.parentOk(a, "mkdir");
        this.nodes.set(a, { kind: "dir", mtime: new Date() });
        return undefined;
    };

    mkdtempSync = (prefix: unknown) => {
        const p = `${this.abs(prefix)}${(++this.tmpCount).toString(36).padStart(6, "0")}`;
        this.mkdirSync(p, { recursive: true });
        return p;
    };

    private children(a: string): string[] {
        const prefix = a === "/" ? "/" : `${a}/`;
        const out: string[] = [];
        for (const p of this.nodes.keys()) {
            if (p !== a && p.startsWith(prefix) && !p.slice(prefix.length).includes("/")) out.push(p.slice(prefix.length));
        }
        return out.sort();
    }

    readdirSync = (p: unknown, opts?: { withFileTypes?: boolean } | string) => {
        const a = this.abs(p);
        const n = this.nodes.get(a);
        if (!n) throw new FsError("ENOENT", "scandir", a);
        if (n.kind !== "dir") throw new FsError("ENOTDIR", "scandir", a);
        const names = this.children(a);
        if (typeof opts === "object" && opts && opts.withFileTypes) {
            return names.map((name) => new Dirent(name, this.nodes.get(path.join(a, name))!.kind === "dir", a));
        }
        return names;
    };

    unlinkSync = (p: unknown) => {
        const a = this.abs(p);
        const n = this.nodes.get(a);
        if (!n) throw new FsError("ENOENT", "unlink", a);
        if (n.kind === "dir") throw new FsError("EISDIR", "unlink", a);
        this.nodes.delete(a);
    };

    rmdirSync = (p: unknown, opts?: { recursive?: boolean }) => {
        const a = this.abs(p);
        const n = this.nodes.get(a);
        if (!n) throw new FsError("ENOENT", "rmdir", a);
        if (n.kind !== "dir") throw new FsError("ENOTDIR", "rmdir", a);
        if (this.children(a).length && !(opts && opts.recursive)) throw new FsError("ENOTEMPTY", "rmdir", a);
        this.removeTree(a);
    };

    rmSync = (p: unknown, opts?: { recursive?: boolean; force?: boolean }) => {
        const a = this.abs(p);
        const n = this.nodes.get(a);
        if (!n) {
            if (opts && opts.force) return;
            throw new FsError("ENOENT", "rm", a);
        }
        if (n.kind === "dir" && !(opts && opts.recursive)) throw new FsError("EISDIR", "rm", a);
        this.removeTree(a);
    };

    private removeTree(a: string) {
        for (const p of [...this.nodes.keys()]) if (p === a || p.startsWith(`${a}/`)) this.nodes.delete(p);
    }

    renameSync = (from: unknown, to: unknown) => {
        const a = this.abs(from);
        const b = this.abs(to);
        if (!this.nodes.has(a)) throw new FsError("ENOENT", "rename", a);
        this.parentOk(b, "rename");
        const moved: [string, Node][] = [];
        for (const [p, n] of this.nodes) if (p === a || p.startsWith(`${a}/`)) moved.push([p, n]);
        this.removeTree(b);
        for (const [p] of moved) this.nodes.delete(p);
        for (const [p, n] of moved) this.nodes.set(b + p.slice(a.length), n);
    };

    copyFileSync = (from: unknown, to: unknown, mode?: number) => {
        const a = this.abs(from);
        const b = this.abs(to);
        const n = this.nodes.get(a);
        if (!n) throw new FsError("ENOENT", "copyfile", a);
        if (n.kind === "dir") throw new FsError("EISDIR", "copyfile", a);
        // COPYFILE_EXCL
        if (mode && mode & 1 && this.nodes.has(b)) throw new FsError("EEXIST", "copyfile", b);
        this.parentOk(b, "copyfile");
        this.nodes.set(b, { kind: "file", data: new Uint8Array(n.data), mtime: new Date() });
    };

    cpSync = (from: unknown, to: unknown, opts?: { recursive?: boolean; force?: boolean; errorOnExist?: boolean }) => {
        const a = this.abs(from);
        const b = this.abs(to);
        const n = this.nodes.get(a);
        if (!n) throw new FsError("ENOENT", "cp", a);
        const force = !opts || opts.force !== false;
        if (n.kind === "file") {
            if (this.nodes.has(b) && !force) {
                if (opts && opts.errorOnExist) throw new FsError("EEXIST", "cp", b);
                return;
            }
            this.mkdirSync(path.dirname(b), { recursive: true });
            this.nodes.set(b, { kind: "file", data: new Uint8Array(n.data), mtime: new Date() });
            return;
        }
        if (!(opts && opts.recursive)) throw new FsError("EISDIR", "cp", a);
        this.mkdirSync(b, { recursive: true });
        for (const name of this.children(a)) this.cpSync(path.join(a, name), path.join(b, name), opts);
    };

    chmodSync = () => {};
    utimesSync = () => {};

    private fdFile(fd: number, syscall: string): string {
        const f = this.fds.get(fd);
        if (!f) throw new FsError("EBADF", syscall, String(fd));
        return f.file;
    }

    openSync = (p: unknown, flags: string | number = "r") => {
        const a = this.abs(p);
        const f = String(flags);
        if (f.startsWith("r")) {
            if (!this.nodes.has(a)) throw new FsError("ENOENT", "open", a);
        } else if (f.startsWith("w") || !this.nodes.has(a)) {
            this.writeFileSync(a, new Uint8Array(0));
        }
        const fd = this.nextFd++;
        this.fds.set(fd, { file: a, append: f.startsWith("a") });
        return fd;
    };

    writeSync = (fd: number, data: unknown, ...rest: unknown[]) => {
        const file = this.fdFile(fd, "write");
        const bytes = toBytes(data, typeof rest[1] === "string" ? rest[1] : undefined);
        this.writeFileSync(file, bytes, { flag: "a" });
        return bytes.length;
    };

    fstatSync = (fd: number) => this.statSync(this.fdFile(fd, "fstat"));

    closeSync = (fd: number) => {
        this.fdFile(fd, "close");
        this.fds.delete(fd);
    };

    realpathSync = Object.assign((p: unknown) => this.abs(p), { native: (p: unknown) => this.abs(p) });

    accessSync = (p: unknown) => {
        const a = this.abs(p);
        if (!this.nodes.has(a)) throw new FsError("ENOENT", "access", a);
    };

    watch = () => ({ close() {} });

    constants = { F_OK: 0, R_OK: 4, W_OK: 2, X_OK: 1, COPYFILE_EXCL: 1 };
}
