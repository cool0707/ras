// Completion of RAS in the playground's editor: the compiler's (crates/ras-wasm `Analysis`), which
// is what the language server offers in VS Code (spec §16.2) — the members of the type before a
// `.`, the names in scope, the exports of modules not imported yet (accepting one adds the
// `import`), and the names of the arguments of a call.

import { type Completion, type CompletionContext, type CompletionResult, type CompletionSource, pickedCompletion } from "@codemirror/autocomplete";
import type { EditorView } from "@codemirror/view";
import type { Text } from "@codemirror/state";

interface Range {
    line: number;
    col: number;
    endLine: number;
    endCol: number;
}

/** An item of `Analysis.completions`. */
export interface Item {
    label: string;
    /** LSP `CompletionItemKind`. */
    kind: number;
    detail: string;
    /** Markdown. */
    doc: string | null;
    /** The group: smaller first. */
    sort: number;
    /** Inserted instead of the label. */
    insert: string | null;
    /** What the item replaces, when that is not the word at the cursor. */
    range?: Range;
    /** The edit that imports the name. */
    import?: Range & { text: string };
}

/** The items at a position of the file being edited (line and column from 0); `namedOnly`:
 * only the names of the arguments of the call. */
export type Ask = (line: number, col: number, namedOnly: boolean) => Item[];

/** LSP `CompletionItemKind` → the icon of CodeMirror. */
const TYPES: Record<number, string> = {
    2: "method",
    3: "function",
    5: "property",
    6: "variable",
    7: "class",
    8: "interface",
    9: "namespace",
    10: "property",
    13: "enum",
    14: "keyword",
    20: "enum",
    21: "constant",
};

const WORD = /[\p{L}\p{N}_$]*$/u;

export function offset(doc: Text, line: number, col: number): number {
    if (line + 1 > doc.lines) return doc.length;
    const l = doc.line(line + 1);
    return Math.min(l.from + col, l.to);
}

/** Markdown of a description (paragraphs, fenced code, `code`, **bold**, *italic*) as elements. */
export function renderDoc(markdown: string, document: Document): HTMLElement {
    const root = document.createElement("div");
    root.className = "doc";
    const inline = (parent: HTMLElement, text: string) => {
        for (const part of text.split(/(`[^`]+`|\*\*[^*]+\*\*|\*[^*\s][^*]*\*)/)) {
            if (!part) continue;
            const tag = part.startsWith("`") ? "code" : part.startsWith("**") ? "strong" : part.startsWith("*") ? "em" : "";
            if (!tag) {
                parent.append(part);
                continue;
            }
            const el = document.createElement(tag);
            const marks = tag === "strong" ? 2 : 1;
            el.textContent = part.slice(marks, -marks);
            parent.append(el);
        }
    };
    const blocks = markdown.split(/```[^\n]*\n([\s\S]*?)\n?```/);
    blocks.forEach((block, i) => {
        if (i % 2) {
            const pre = document.createElement("pre");
            pre.textContent = block;
            root.append(pre);
            return;
        }
        for (const para of block.split(/\n{2,}/)) {
            if (!para.trim()) continue;
            const p = document.createElement("p");
            inline(p, para.trim());
            root.append(p);
        }
    });
    return root;
}

function option(item: Item, doc: Text): Completion {
    const text = item.insert ?? item.label;
    const c: Completion = {
        label: item.label,
        type: TYPES[item.kind],
        // The groups only order the items that match the word equally well.
        boost: (4 - Math.min(item.sort, 4)) * 20,
    };
    if (item.detail) c.detail = item.detail;
    if (item.doc) {
        const markdown = item.doc;
        c.info = () => renderDoc(markdown, document);
    }
    if (item.range || item.import) {
        // Positions of the text the items were computed for: the items are asked for again
        // when it changes (see `validFor`), except for the word at the cursor.
        const range = item.range && { from: offset(doc, item.range.line, item.range.col) };
        const imp = item.import && {
            from: offset(doc, item.import.line, item.import.col),
            to: offset(doc, item.import.endLine, item.import.endCol),
            insert: item.import.text,
        };
        c.apply = (view: EditorView, picked: Completion, from: number, to: number) => {
            const start = range ? range.from : from;
            const changes = [{ from: start, to, insert: text }, ...(imp ? [imp] : [])];
            const set = view.state.changes(changes);
            view.dispatch({ changes: set, selection: { anchor: set.mapPos(to, 1) }, userEvent: "input.complete", annotations: pickedCompletion.of(picked) });
        };
    } else if (text !== item.label) {
        c.apply = text;
    }
    return c;
}

/** The completion source of a .ras file, from the compiler's answers. */
export function rasCompletions(ask: Ask): CompletionSource {
    return (context: CompletionContext): CompletionResult | null => {
        const doc = context.state.doc;
        const line = doc.lineAt(context.pos);
        const before = line.text.slice(0, context.pos - line.from);
        const word = WORD.exec(before)![0];
        const from = context.pos - word.length;
        const prev = before[before.length - word.length - 1] ?? "";
        // A number is not a name, and `1.` is not a member access.
        if (/^[0-9]/.test(word)) return null;
        // Typing `(` or `,` only offers the names of the arguments (spec §3.3).
        const namedOnly = !word && !context.explicit && (prev === "(" || prev === ",");
        if (!word && !context.explicit && prev !== "." && !namedOnly) return null;
        if (prev === "." && /(?:^|[^\p{L}\p{N}_$])[0-9]+\.$/u.test(before.slice(0, from))) return null;
        let items: Item[];
        try {
            items = ask(line.number - 1, context.pos - line.from, namedOnly);
        } catch {
            return null;
        }
        if (!items.length) return null;
        const start = items.reduce((p, i) => (i.range ? Math.min(p, offset(doc, i.range.line, i.range.col)) : p), from);
        return {
            from: start,
            options: items.map((i) => option(i, doc)),
            // The names of the arguments give way to everything at the next keystroke.
            validFor: namedOnly ? undefined : start === from ? /^[\p{L}\p{N}_$]*$/u : /^[\p{L}\p{N}_$.]*$/u,
        };
    };
}
