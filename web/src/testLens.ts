// "テストを生成" above the functions a test can be generated for (spec §14.12.10): the code
// lens of the VS Code extension, as a line above the function in the editor.

import { type Extension, RangeSetBuilder, StateEffect, StateField } from "@codemirror/state";
import { Decoration, type DecorationSet, EditorView, WidgetType } from "@codemirror/view";

/** A function of the file: its name and the line (from 0) of its declaration. */
export interface TestLens {
    name: string;
    line: number;
}

const setLenses = StateEffect.define<TestLens[]>();

class LensWidget extends WidgetType {
    constructor(
        readonly lens: TestLens,
        readonly indent: string,
        readonly onClick: (name: string) => void,
    ) {
        super();
    }

    eq(other: LensWidget): boolean {
        return other.lens.name === this.lens.name && other.indent === this.indent;
    }

    toDOM(): HTMLElement {
        const div = document.createElement("div");
        div.className = "cm-test-lens";
        div.style.paddingLeft = `${this.indent.length}ch`;
        const a = document.createElement("button");
        a.type = "button";
        a.title = `${this.lens.name} のテストの雛形を test/ に追加して開きます`;
        const icon = document.createElement("i");
        icon.className = "codicon codicon-beaker";
        icon.setAttribute("aria-hidden", "true");
        a.append(icon, " テストを生成");
        a.onmousedown = (e) => e.preventDefault();
        a.onclick = (e) => {
            e.preventDefault();
            this.onClick(this.lens.name);
        };
        div.append(a);
        return div;
    }

    ignoreEvent(): boolean {
        return true;
    }
}

function build(state: EditorView["state"], lenses: TestLens[], onClick: (name: string) => void): DecorationSet {
    const b = new RangeSetBuilder<Decoration>();
    const doc = state.doc;
    for (const l of [...lenses].sort((a, b) => a.line - b.line)) {
        if (l.line + 1 > doc.lines) continue;
        const line = doc.line(l.line + 1);
        const indent = /^\s*/.exec(line.text)?.[0] ?? "";
        b.add(line.from, line.from, Decoration.widget({ widget: new LensWidget(l, indent, onClick), block: true, side: -1 }));
    }
    return b.finish();
}

/** The lenses of the editor; `onClick` is called with the name of the function. */
export function testLenses(onClick: (name: string) => void): Extension {
    const field = StateField.define<DecorationSet>({
        create: () => Decoration.none,
        update(value, tr) {
            for (const e of tr.effects) if (e.is(setLenses)) return build(tr.state, e.value, onClick);
            return value.map(tr.changes);
        },
        provide: (f) => EditorView.decorations.from(f),
    });
    return [
        field,
        EditorView.baseTheme({
            ".cm-test-lens": { fontSize: "0.85em", lineHeight: "1.6", color: "var(--muted)", userSelect: "none" },
            ".cm-test-lens button": {
                font: "inherit",
                color: "inherit",
                background: "none",
                border: "none",
                padding: "0",
                cursor: "pointer",
            },
            ".cm-test-lens button:hover": { color: "var(--accent)", textDecoration: "underline" },
            ".cm-test-lens .codicon": { fontSize: "inherit", verticalAlign: "-1px" },
        }),
    ];
}

/** Shows the lenses of the file in the editor (replacing the previous ones). */
export function showTestLenses(view: EditorView, lenses: TestLens[]): void {
    view.dispatch({ effects: setLenses.of(lenses) });
}
