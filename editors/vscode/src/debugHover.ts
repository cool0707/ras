// Hovers while debugging (spec §14.10). The debug build reads members through the runtime
// (`sheet.name` becomes `$rt.get(sheet, "Name")`), so VS Code's default hover — evaluating the
// source text `sheet.name` as JavaScript — finds nothing on a host object and shows `undefined`.
// The hovered member chain is rewritten into a call of the runtime's `$rasDebugRead`, which reads
// it the way the program does.

const IDENT = /[A-Za-z_$][\w$]*/y;

export interface HoverExpression {
	/** Columns of the hovered chain in the line (end exclusive). */
	start: number;
	end: number;
	expression: string;
}

/**
 * The expression to evaluate for a hover at `column` of `line`, or `undefined` to leave it to
 * VS Code. Only plain `a.b.c` chains are rewritten: a name that is called (`ws.delete()`), or a
 * chain through a call or index (`ws.cells(1, 1).value`), is not evaluated, as reading it could
 * change the workbook.
 */
export function hoverExpression(line: string, column: number): HoverExpression | undefined {
	const word = wordAt(line, column);
	if (word === undefined || inStringOrComment(line, word.start)) {
		return undefined;
	}
	if (/^\s*\(/.test(line.slice(word.end))) {
		return undefined;
	}
	const names = [line.slice(word.start, word.end)];
	let start = word.start;
	for (;;) {
		const before = /([A-Za-z_$][\w$]*)\s*\??\.\s*$/.exec(line.slice(0, start));
		if (before === null) {
			break;
		}
		start = before.index;
		names.unshift(before[1]);
	}
	// `a.b` where `a` itself follows a call, an index or a literal.
	if (/[)\].'"`\d]\s*$/.test(line.slice(0, start))) {
		return undefined;
	}
	const [root, ...path] = names;
	if (/^\d/.test(root)) {
		return undefined;
	}
	return { start, end: word.end, expression: readExpression(root, path) };
}

function readExpression(root: string, path: string[]): string {
	// Names the compiler had to change (`this` of an extension method, JS reserved words) are
	// tried first; the runtime falls back to the next candidate, then to a host global.
	const candidates = root === "this" ? ["$this", "this"] : [`${root}$`, root];
	const roots = candidates.map((name) => `() => ${name}`).join(", ");
	return `$rasDebugRead([${roots}], ${JSON.stringify(root)}, ${JSON.stringify(path)})`;
}

function wordAt(line: string, column: number): { start: number; end: number } | undefined {
	for (let i = 0; i < line.length; ) {
		IDENT.lastIndex = i;
		const m = IDENT.exec(line);
		if (m === null) {
			i++;
			continue;
		}
		const end = i + m[0].length;
		// A number (`1e5`) or the tail of one is not a name.
		if (i > 0 && /[\w$]/.test(line[i - 1])) {
			i = end;
			continue;
		}
		if (column >= i && column <= end) {
			return { start: i, end };
		}
		i = end;
	}
	return undefined;
}

/** Whether `column` is inside a string literal or after `//` (single-line view only). */
function inStringOrComment(line: string, column: number): boolean {
	let quote: string | undefined;
	for (let i = 0; i < column; i++) {
		const c = line[i];
		if (quote !== undefined) {
			if (c === "\\") {
				i++;
			} else if (c === quote) {
				quote = undefined;
			}
		} else if (c === '"' || c === "'" || c === "`") {
			quote = c;
		} else if (c === "/" && line[i + 1] === "/") {
			return true;
		}
	}
	return quote !== undefined;
}
