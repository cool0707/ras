// Hovers and parameter information of RAS in the playground's editor, from the compiler
// (crates/ras-wasm `Analysis`), as the language server gives them in VS Code (spec §16.2):
// pointing at a name shows what it is, and in the arguments of a call a tooltip above it shows
// the signature with the parameter being written.

import { type EditorState, type Extension, Prec, StateEffect, StateField } from "@codemirror/state";
import { type EditorView, hoverTooltip, keymap, showTooltip, type Tooltip, ViewPlugin, type ViewUpdate } from "@codemirror/view";
import { offset, renderDoc } from "./complete";

/** `Analysis.hover`: the description (Markdown) of a name and its range. */
export interface HoverInfo {
    markdown: string;
    line: number;
    col: number;
    endLine: number;
    endCol: number;
}

/** A signature of `Analysis.signatureHelp`. */
export interface Signature {
    label: string;
    doc: string | null;
    /** Ranges of the label, with the `@param` descriptions. */
    params: { from: number; to: number; doc: string | null }[];
    /** The parameter being written. */
    active: number | null;
}

/** `Analysis.signatureHelp`: the signatures of a call and the one that fits its arguments. */
export interface SignatureInfo {
    signatures: Signature[];
    active: number;
}

/** The compiler's answers at a position of the file being edited (line and column from 0). */
export interface Asks {
    hover(line: number, col: number): HoverInfo | null;
    signature(line: number, col: number): SignatureInfo | null;
}

function lineCol(state: EditorState, pos: number): [number, number] {
    const line = state.doc.lineAt(pos);
    return [line.number - 1, pos - line.from];
}

/** How far back a call's `(` is looked for. */
const LOOK_BACK = 4000;

/**
 * The `(` of the innermost argument list `pos` is in, or -1: a look at the text alone, so that
 * the compiler is only asked inside a call.
 */
export function openParen(state: EditorState, pos: number): number {
    const start = Math.max(0, pos - LOOK_BACK);
    const text = state.doc.sliceString(start, pos);
    const open: number[] = [];
    for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (c === '"' || c === "'") {
            // A string (on one line); one that is still open hides nothing.
            let j = i + 1;
            while (j < text.length && text[j] !== c && text[j] !== "\n") j += text[j] === "\\" ? 2 : 1;
            if (j >= text.length || text[j] === "\n") continue;
            i = j;
        } else if (c === "/" && text[i + 1] === "/") {
            const end = text.indexOf("\n", i);
            if (end < 0) break;
            i = end;
        } else if (c === "(") open.push(i);
        else if (c === ")") open.pop();
    }
    return open.length ? start + open[open.length - 1] : -1;
}

/** The tooltip of a call: its signatures, the parameter being written in bold. */
export function renderSignatures(info: SignatureInfo, document: Document): HTMLElement {
    const root = document.createElement("div");
    root.className = "cm-signature";
    info.signatures.forEach((sig, i) => {
        const current = i === info.active;
        const line = document.createElement("div");
        line.className = current ? "sig active" : "sig";
        const p = current && sig.active !== null ? sig.params[sig.active] : undefined;
        if (p) {
            const b = document.createElement("strong");
            b.textContent = sig.label.slice(p.from, p.to);
            line.append(sig.label.slice(0, p.from), b, sig.label.slice(p.to));
        } else {
            line.textContent = sig.label;
        }
        root.append(line);
        if (!current) return;
        if (p?.doc) {
            const doc = renderDoc(p.doc, document);
            const name = document.createElement("code");
            name.textContent = sig.label.slice(p.from, p.to).replace(/[?:=].*$/s, "").trim();
            doc.firstElementChild?.prepend(name, " — ");
            root.append(doc);
        }
        if (sig.doc) root.append(renderDoc(sig.doc, document));
    });
    return root;
}

const setSignature = StateEffect.define<Tooltip | null>();

const signatureField = StateField.define<Tooltip | null>({
    create: () => null,
    update(value, tr) {
        for (const e of tr.effects) if (e.is(setSignature)) return e.value;
        return value;
    },
    provide: (f) => showTooltip.from(f),
});

/** What opens the parameter information when typed: `(` (which the editor closes: `()`) or `,`. */
const TRIGGER = /^\(|[(,]$/;

function parameterHints(asks: Asks): Extension {
    const plugin = ViewPlugin.fromClass(
        class {
            /** Whether the tooltip follows the cursor (since a `(` or `,` was typed, until the
             * cursor leaves the call or Escape). */
            open = false;
            timer = 0;

            constructor(readonly view: EditorView) {}

            update(u: ViewUpdate) {
                let typed = false;
                for (const tr of u.transactions) {
                    if (!tr.docChanged || !(tr.isUserEvent("input.type") || tr.isUserEvent("input.complete"))) continue;
                    tr.changes.iterChanges((_a, _b, _c, _d, inserted) => {
                        if (TRIGGER.test(inserted.toString())) typed = true;
                    });
                }
                if (typed) this.open = true;
                if (u.focusChanged && !u.view.hasFocus) this.close();
                else if (this.open && (u.docChanged || u.selectionSet)) this.schedule();
            }

            schedule() {
                clearTimeout(this.timer);
                this.timer = window.setTimeout(() => this.refresh(), 60);
            }

            show() {
                this.open = true;
                this.schedule();
            }

            close(): boolean {
                clearTimeout(this.timer);
                const shown = this.view.state.field(signatureField) !== null;
                this.open = false;
                // Not during an update of the view (`update` calls this when focus is lost).
                if (shown) queueMicrotask(() => this.view.dispatch({ effects: setSignature.of(null) }));
                return shown;
            }

            refresh() {
                if (!this.open) return;
                const state = this.view.state;
                const pos = state.selection.main.head;
                const paren = state.selection.main.empty ? openParen(state, pos) : -1;
                let info: SignatureInfo | null = null;
                if (paren >= 0) {
                    try {
                        info = asks.signature(...lineCol(state, pos));
                    } catch {
                        info = null;
                    }
                }
                if (!info || !info.signatures.length) {
                    this.open = false;
                    if (state.field(signatureField)) this.view.dispatch({ effects: setSignature.of(null) });
                    return;
                }
                const found = info;
                this.view.dispatch({
                    effects: setSignature.of({
                        pos: paren,
                        above: true,
                        create: () => ({ dom: renderSignatures(found, document) }),
                    }),
                });
            }

            destroy() {
                clearTimeout(this.timer);
            }
        },
    );
    return [
        signatureField,
        plugin,
        Prec.high(
            keymap.of([
                { key: "Mod-Shift-Space", run: (view) => (view.plugin(plugin)?.show(), true) },
                { key: "Escape", run: (view) => view.plugin(plugin)?.close() ?? false },
            ]),
        ),
    ];
}

/** The hovers and the parameter information of a .ras file, from the compiler's answers. */
export function rasHints(asks: Asks): Extension {
    const hover = hoverTooltip((view, pos) => {
        let info: HoverInfo | null;
        try {
            info = asks.hover(...lineCol(view.state, pos));
        } catch {
            return null;
        }
        if (!info) return null;
        const found = info;
        const doc = view.state.doc;
        return {
            pos: offset(doc, found.line, found.col),
            end: offset(doc, found.endLine, found.endCol),
            above: true,
            create: () => {
                const dom = renderDoc(found.markdown, document);
                dom.classList.add("cm-hover");
                return { dom };
            },
        };
    });
    return [hover, parameterHints(asks)];
}
