// The worker the playground runs programs in, so that a long or endless run does not stop the
// page (it is terminated instead). It is given the compiled WebAssembly module and a run.
//
// Dialogs: when the page is cross-origin isolated (see sw.js), the worker waits on a shared
// buffer while the page shows the dialog; otherwise they get the answers `ras run` gives.

import init, { Book } from "../pkg/ras_wasm.js";
import { runJob, ServeSession, type Ask, type Dialog, type FormChannel, type RunJob, type ServeReply, type ServeSnapshot, type WasmBookClass } from "./runtime/run";

/** `shared`: the dialogs' answers; `forms`: what the user does on the forms (both as `asker` lays them out). */
export type ToWorker =
    | { type: "run"; module: WebAssembly.Module; job: RunJob; shared: SharedArrayBuffer | null; forms: SharedArrayBuffer | null }
    /** The run mode of the book tab: the program started (`ras-main.js --serve`), then its `start` request. */
    | { type: "serve"; module: WebAssembly.Module; job: RunJob; shared: SharedArrayBuffer | null; forms: SharedArrayBuffer | null }
    /** What was done on the sheet in run mode (`$rt.serve`'s requests). */
    | { type: "request"; id: number; request: Record<string, unknown> };
export type FromWorker =
    | { type: "output"; stream: "stdout" | "stderr"; text: string }
    | { type: "dialog"; dialog: Dialog }
    /** A form's state for the page to draw (spec §14.10.7). */
    | { type: "form"; message: unknown }
    /** The run waits for what the user does next on the forms. */
    | { type: "formWait" }
    | { type: "formPoll" }
    | { type: "done"; result: ReturnType<typeof runJob>; ms: number }
    | { type: "failed"; message: string }
    /** The answer of a request of the run mode (`id`: the request's; 0 for `start`). */
    | { type: "reply"; id: number; reply: ServeReply; snapshot: ServeSnapshot | null; exited?: boolean };

const post = (m: FromWorker, transfer: Transferable[] = []) => (self as unknown as Worker).postMessage(m, transfer);

/** The shared buffer: [0] the state (0 waiting, 1 answered), [1] the length of the answer's JSON, then its bytes. */
function asker(shared: SharedArrayBuffer): Ask {
    const state = new Int32Array(shared, 0, 2);
    const bytes = new Uint8Array(shared, 8);
    return (dialog) => {
        Atomics.store(state, 0, 0);
        post({ type: "dialog", dialog });
        Atomics.wait(state, 0, 0);
        const text = new TextDecoder().decode(bytes.slice(0, state[1]));
        const answer = JSON.parse(text) as { value: number | string | false | null };
        return answer.value === null ? undefined : answer.value;
    };
}

/** The forms are drawn by the page: their states are posted, the user's actions waited for. */
function formChannel(shared: SharedArrayBuffer): FormChannel {
    const state = new Int32Array(shared, 0, 2);
    const bytes = new Uint8Array(shared, 8);
    return {
        send: (message) => post({ type: "form", message }),
        receive: () => {
            Atomics.store(state, 0, 0);
            post({ type: "formWait" });
            Atomics.wait(state, 0, 0);
            return JSON.parse(new TextDecoder().decode(bytes.slice(0, state[1])));
        },
        // What the user has done, without waiting for more (DoEvents with a modeless form shown).
        poll: () => {
            Atomics.store(state, 0, 0);
            post({ type: "formPoll" });
            Atomics.wait(state, 0, 0);
            return JSON.parse(new TextDecoder().decode(bytes.slice(0, state[1])));
        },
    };
}

let session: ServeSession | null = null;

function reply(id: number, request: Record<string, unknown>) {
    const r = session!.request(request);
    post({ type: "reply", id, ...r, exited: session!.exited !== null }, r.snapshot ? [r.snapshot.bytes.buffer] : []);
}

self.onmessage = async (e: MessageEvent<ToWorker>) => {
    const m = e.data;
    if (m.type === "request") {
        try {
            reply(m.id, m.request);
        } catch (err) {
            post({ type: "reply", id: m.id, reply: { id: 0, ok: false, error: { code: 0, description: String((err && (err as Error).message) || err), file: null, line: 0 } }, snapshot: null });
        }
        return;
    }
    if (m.type === "serve") {
        try {
            await init({ module_or_path: m.module });
            const output = (stream: "stdout" | "stderr", text: string) => post({ type: "output", stream, text });
            session = new ServeSession(Book as unknown as WasmBookClass, m.job, output, m.shared ? asker(m.shared) : undefined, m.forms ? formChannel(m.forms) : undefined);
            reply(0, { op: "start" });
        } catch (err) {
            post({ type: "failed", message: String((err && (err as Error).stack) || err) });
        }
        return;
    }
    if (m.type !== "run") return;
    try {
        await init({ module_or_path: m.module });
        const started = performance.now();
        const result = runJob(
            Book as unknown as WasmBookClass,
            m.job,
            (stream, text) => post({ type: "output", stream, text }),
            m.shared ? asker(m.shared) : undefined,
            m.forms ? formChannel(m.forms) : undefined,
        );
        post({ type: "done", result, ms: performance.now() - started }, result.saved.map((s) => s.data.buffer));
    } catch (err) {
        post({ type: "failed", message: String((err && (err as Error).stack) || err) });
    }
};
