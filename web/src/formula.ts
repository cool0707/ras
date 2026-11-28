// Excel formulas in `fx`…`` (spec §2.7) in the playground's editor: the formula is parsed by a
// language of its own nested in the template (its functions, references, sheets, strings… are
// coloured), and completion there offers Excel's functions instead of RAS's names. The functions
// are those the language server offers (crates/ras-compiler/stdlib/formula_functions.json).

import type { Completion, CompletionContext, CompletionResult, CompletionSource } from "@codemirror/autocomplete";
import { LanguageSupport, type LRLanguage, StreamLanguage, type StringStream } from "@codemirror/language";
import { javascript } from "@codemirror/lang-javascript";
import { tags } from "@lezer/highlight";
import { parseMixed, type SyntaxNodeRef } from "@lezer/common";
import functionList from "../../crates/ras-compiler/stdlib/formula_functions.json";

/** [name, signature, description] of every function, by name. */
export const FUNCTIONS = functionList as [string, string, string][];

const WORD = /[\p{L}\p{N}_.]/u;

interface ScanState {
    inString: boolean;
}

/** The tokens of a formula (a line of a template, without its `${…}`). */
function token(stream: StringStream, state: ScanState): string | null {
    if (state.inString) {
        while (!stream.eol()) {
            if (stream.next() === '"') {
                if (stream.eat('"')) continue;
                state.inString = false;
                break;
            }
        }
        return "string";
    }
    if (stream.eatSpace()) return null;
    if (stream.eat('"')) {
        state.inString = true;
        return token(stream, state) ?? "string";
    }
    // 'Sheet name'! / Sheet1!
    if (stream.match(/^'(?:[^']|'')*'!/) || stream.match(/^(?:\[[^\]]+\])?[\p{L}_][\p{L}\p{N}_.]*!/u)) return "sheet";
    if (stream.match(/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)/)) return "errorValue";
    // SUM( / STDEV.S(
    if (stream.match(/^(?:_xlfn\.)?[\p{L}_][\p{L}\p{N}_.]*(?=\s*\()/iu)) return "function";
    if (stream.match(/^(?:TRUE|FALSE)(?![\p{L}\p{N}_.(])/iu)) return "bool";
    // A1, $B$2, A:C, R[-1]C2
    if (stream.match(/^(?:\$?[A-Z]{1,3}\$?[0-9]+|R(?:\[-?[0-9]+\]|[0-9]+)?C(?:\[-?[0-9]+\]|[0-9]+)?)(?![\p{L}\p{N}_.(])/iu)) return "reference";
    if (stream.match(/^\$?[A-Z]{1,3}(?=\s*:\s*\$?[A-Z]{1,3}(?![\p{L}\p{N}_.]))/iu)) return "reference";
    if (stream.string[stream.pos - 1] === ":" && stream.match(/^\$?[A-Z]{1,3}(?![\p{L}\p{N}_.(])/iu)) return "reference";
    if (stream.match(/^(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/)) return "number";
    if (stream.match(/^[\p{L}_][\p{L}\p{N}_.]*/u)) return "name";
    if (stream.match(/^(?:<>|<=|>=|[-+*/^%&=<>:@#])/)) return "operator";
    if (stream.match(/^[(){}[\]]/)) return "bracket";
    if (stream.match(/^[,;]/)) return "separator";
    stream.next();
    return null;
}

/** The completion at the cursor in a formula: the functions that begin with the word written. */
export function formulaCompletions(context: CompletionContext): CompletionResult | null {
    const line = context.state.doc.lineAt(context.pos);
    const before = line.text.slice(0, context.pos - line.from);
    // In a string of the formula: nothing (count the quotes since the template began).
    const tick = before.lastIndexOf("`");
    if (((before.slice(tick + 1).match(/"/g) || []).length & 1) === 1) return null;
    let start = before.length;
    while (start > 0 && WORD.test(before[start - 1])) start--;
    const word = before.slice(start);
    const prev = before[start - 1] || "";
    if (/^[0-9.]/.test(word) || /^\$?[A-Z]{1,3}[0-9]+$/i.test(word) || "!$:'".includes(prev) && prev !== "") return null;
    if (!word && !context.explicit) return null;
    const options: Completion[] = FUNCTIONS.map(([name, signature, description]) => ({
        label: name,
        type: "function",
        detail: signature,
        info: description,
        apply: `${name}(`,
    }));
    return { from: line.from + start, options, validFor: /^[\p{L}\p{N}_.]*$/u };
}

/** The language of a formula, with its completion. */
export const formulaLanguage = StreamLanguage.define<ScanState>({
    name: "excel-formula",
    startState: () => ({ inString: false }),
    copyState: (s) => ({ ...s }),
    token,
    languageData: { autocomplete: formulaCompletions },
    tokenTable: {
        function: tags.function(tags.variableName),
        reference: tags.special(tags.variableName),
        sheet: tags.namespace,
        errorValue: tags.atom,
        bool: tags.bool,
        name: tags.variableName,
        separator: tags.separator,
    },
});

/** Whether `node` (a TemplateString) is the formula of a `fx`…`` / `obj.fx`…``. */
function isFormula(node: SyntaxNodeRef, text: (from: number, to: number) => string): boolean {
    const parent = node.node.parent;
    if (!parent || parent.name !== "TaggedTemplateExpression") return false;
    const tag = parent.firstChild;
    if (!tag || tag.to > node.from) return false;
    if (tag.name === "VariableName") return text(tag.from, tag.to) === "fx";
    if (tag.name === "MemberExpression") return tag.lastChild?.name === "PropertyName" && text(tag.lastChild.from, tag.lastChild.to) === "fx";
    return false;
}

/** TypeScript (RAS is close enough to it for its highlighting and brackets) with the formulas
 * of `fx` templates nested. `complete` is the completion outside the formulas (the compiler's,
 * `complete.ts`); without it, TypeScript's. */
export function rasLanguage(complete?: CompletionSource): LanguageSupport {
    const ts = javascript({ typescript: true });
    const language = (ts.language as LRLanguage).configure({
        wrap: parseMixed((node, input) => {
            if (node.name !== "TemplateString" || !isFormula(node, (a, b) => input.read(a, b))) return null;
            // The text between the back-ticks, without the `${…}`.
            const overlay: { from: number; to: number }[] = [];
            let from = node.from + 1;
            for (let child = node.node.firstChild; child; child = child.nextSibling) {
                if (child.name !== "Interpolation") continue;
                if (child.from > from) overlay.push({ from, to: child.from });
                from = child.to;
            }
            const end = input.read(node.to - 1, node.to) === "`" ? node.to - 1 : node.to;
            if (end > from) overlay.push({ from, to: end });
            return overlay.length ? { parser: formulaLanguage.parser, overlay } : null;
        }),
    });
    return new LanguageSupport(language, complete ? language.data.of({ autocomplete: complete }) : ts.support);
}
