// The user forms of a run in the playground (spec §14.10.7): the page is the host that draws
// them, as the VS Code extension's FormHost (editors/vscode/src/formRunner/host.ts) is. The run
// (in the worker) posts the state of each form; each form is drawn by the extension's form
// panel (webview/formRunner) in a frame of its own over the page, and what the user does there
// is handed back to the run, which waits for it on a shared buffer.

import { applyUpdate, initialState, px } from "../../editors/vscode/src/formRunner/protocol";
import type { FormState, PanelMessage, RuntimeMessage, UserAction } from "../../editors/vscode/src/formRunner/protocol";

interface Panel {
    frame: HTMLIFrameElement;
    box: HTMLElement;
    state: FormState;
    ready: boolean;
    pending: PanelMessage[];
}

export class FormPanels {
    private readonly panels = new Map<number, Panel>();
    /** What the user did, not yet taken by the run. */
    private readonly queue: (UserAction & { id: number })[] = [];
    private waiting = false;
    private readonly layer: HTMLElement;

    /** `answer` hands the run what the user did (it is waiting for it). */
    constructor(private readonly answer: (action: unknown) => void) {
        this.layer = document.createElement("div");
        this.layer.className = "form-layer";
        this.layer.hidden = true;
        document.body.append(this.layer);
        window.addEventListener("message", (e) => {
            const panel = [...this.panels.entries()].find(([, p]) => p.frame.contentWindow === e.source);
            if (!panel) return;
            const m = e.data && (e.data as { rasFrame?: { type?: string; action?: UserAction } }).rasFrame;
            if (!m) return;
            const [id, p] = panel;
            if (m.type === "ready") {
                p.ready = true;
                for (const x of p.pending) p.frame.contentWindow?.postMessage(x, "*");
                p.pending = [];
            } else if (m.type === "action" && m.action) {
                this.queue.push({ ...m.action, id });
                this.flush();
            }
        });
    }

    /** A message of the run (`show`, `update`, ...). */
    receive(msg: RuntimeMessage): void {
        switch (msg.t) {
            case "show": {
                const p = this.panels.get(msg.id);
                if (p) {
                    p.state = initialState(msg);
                    this.post(p, { type: "state", state: p.state });
                    this.layout(p);
                } else this.open(msg.id, initialState(msg));
                break;
            }
            case "update": {
                const p = this.panels.get(msg.id);
                if (!p) break;
                p.state = applyUpdate(p.state, msg);
                this.post(p, { type: "state", state: p.state, focus: msg.focus });
                this.layout(p);
                break;
            }
            case "hide":
            case "unload":
                this.close(msg.id);
                break;
            case "error": {
                const p = this.panels.get(msg.id);
                if (p) this.post(p, { type: "error", message: msg.message });
                break;
            }
            case "blocked": {
                const p = this.panels.get(msg.id);
                if (p) this.post(p, { type: "blocked" });
                break;
            }
        }
    }

    /** The run waits for what the user does next. */
    wait(): void {
        this.waiting = true;
        this.flush();
    }

    /** The run has ended: its forms are gone. */
    end(): void {
        for (const id of [...this.panels.keys()]) this.close(id);
        this.queue.length = 0;
        this.waiting = false;
    }

    private flush(): void {
        if (!this.waiting || !this.queue.length) return;
        this.waiting = false;
        this.answer(this.queue.shift());
    }

    private post(p: Panel, message: PanelMessage): void {
        if (p.ready) p.frame.contentWindow?.postMessage(message, "*");
        else if (message.type === "state") p.pending = [...p.pending.filter((m) => m.type !== "state"), message];
        else p.pending.push(message);
    }

    private open(id: number, state: FormState): void {
        const box = document.createElement("div");
        box.className = "form-box";
        const frame = document.createElement("iframe");
        frame.title = `フォーム ${state.name}`;
        frame.src = "runner.html";
        box.append(frame);
        // Forms shown later are on top of the earlier ones, a little to the lower right.
        const n = this.panels.size;
        box.style.transform = `translate(${n * 24}px, ${n * 24}px)`;
        this.layer.append(box);
        this.layer.hidden = false;
        const p: Panel = { frame, box, state, ready: false, pending: [{ type: "state", state }] };
        this.panels.set(id, p);
        this.layout(p);
    }

    /** The frame fits the form: its size (points) with the title bar and the panel's margins. */
    private layout(p: Panel): void {
        const w = parseFloat(px(p.state.form.Width)) || 300;
        const h = parseFloat(px(p.state.form.Height)) || 200;
        p.frame.style.width = `${Math.max(Math.ceil(w) + 40, 380)}px`;
        p.frame.style.height = `${Math.ceil(h) + 130}px`;
        this.layer.classList.toggle("modal", [...this.panels.values()].some((x) => x.state.modal));
    }

    private close(id: number): void {
        const p = this.panels.get(id);
        if (!p) return;
        p.box.remove();
        this.panels.delete(id);
        if (!this.panels.size) this.layer.hidden = true;
        else this.layer.classList.toggle("modal", [...this.panels.values()].some((x) => x.state.modal));
    }
}
