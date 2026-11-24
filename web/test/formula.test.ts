// The formulas of `fx`…`` in the playground's editor (spec §2.7): highlighted as Excel's and
// completed with Excel's functions (`npm test`; this file does not need web/pkg).
import assert from "node:assert/strict";
import { test } from "node:test";
import { CompletionContext, type CompletionResult, type CompletionSource } from "@codemirror/autocomplete";
import { ensureSyntaxTree } from "@codemirror/language";
import { EditorState } from "@codemirror/state";
import { highlightTree, tagHighlighter, tags } from "@lezer/highlight";
import { FUNCTIONS, rasLanguage } from "../src/formula";

const classes = tagHighlighter([
    { tag: tags.function(tags.variableName), class: "fn" },
    { tag: tags.special(tags.variableName), class: "ref" },
    { tag: tags.namespace, class: "sheet" },
    { tag: tags.string, class: "string" },
    { tag: tags.atom, class: "atom" },
]);

/** The class of each highlighted piece of `doc`: "text=class". */
function highlighted(doc: string): string[] {
    const state = EditorState.create({ doc, extensions: [rasLanguage()] });
    const tree = ensureSyntaxTree(state, doc.length, 5000);
    assert.ok(tree);
    const out: string[] = [];
    highlightTree(tree, classes, (from, to, cls) => out.push(`${doc.slice(from, to)}=${cls}`));
    return out;
}

function complete(doc: string, explicit = false): CompletionResult | null {
    const state = EditorState.create({ doc, extensions: [rasLanguage()] });
    ensureSyntaxTree(state, doc.length, 5000);
    const sources = state.languageDataAt<CompletionSource>("autocomplete", doc.length, -1);
    const context = new CompletionContext(state, doc.length, explicit);
    for (const source of sources) {
        if (typeof source !== "function") continue;
        const r = source(context);
        if (r && !(r instanceof Promise) && r.options.length) return r as CompletionResult;
    }
    return null;
}

test("formulas of fx templates are highlighted as Excel's", () => {
    const got = highlighted('const v = ws.fx`SUM(Sheet1!A1:B2, "x", #N/A) + ${n}`;');
    for (const want of ["SUM=fn", "Sheet1!=sheet", "A1=ref", "B2=ref", '"x"=string', "#N/A=atom"]) {
        assert.ok(got.includes(want), `${want} in ${got.join(" ")}`);
    }
    // Any other template is a string.
    const plain = highlighted("const v = tag`SUM(A1)`;");
    assert.ok(!plain.includes("SUM=fn") && !plain.includes("A1=ref"), plain.join(" "));
});

test("completion in a formula offers Excel's functions", () => {
    const r = complete("const v = fx`1 + VLO");
    assert.ok(r);
    assert.equal(r.from, "const v = fx`1 + ".length);
    const vlookup = r.options.find((o) => o.label === "VLOOKUP");
    assert.ok(vlookup);
    assert.equal(vlookup.apply, "VLOOKUP(");
    assert.equal(r.options.length, FUNCTIONS.length);
    // Not in its strings or references.
    assert.equal(complete('const v = fx`COUNTIF(A:A, "VLO'), null);
    assert.equal(complete("const v = fx`Sheet1!B"), null);
    // Outside the formula: RAS (TypeScript) completion, not Excel's.
    const outside = complete("const vlo = 1; vl");
    assert.ok(!outside || !outside.options.some((o) => o.label === "VLOOKUP"));
});
