// The `list` of a ComboBox / ListBox as the text of the Property Grid's multi-line editor
// (one item per line, its columns separated by tabs), and the columns of a list as MSForms
// lays them out (ColumnCount, ColumnWidths, TextColumn).

/** The texts of the columns of an item as a FormSpec writes it: a string, or one per column. */
export function itemColumns(item: unknown): string[] {
	if (Array.isArray(item)) return item.map((column) => (typeof column === "string" ? column : String(column ?? "")));
	return [typeof item === "string" ? item : String(item ?? "")];
}

const isItem = (item: unknown) => typeof item === "string" || (Array.isArray(item) && item.every((column) => typeof column === "string"));

/** Text shown for a `list` value; anything that is not a list of items is empty. */
export function listToText(value: unknown): string {
	return Array.isArray(value) && value.every(isItem) ? value.map((item) => itemColumns(item).join("\t")).join("\n") : "";
}

/**
 * Items entered in the editor: a line with tabs is an item of several columns. Blank lines
 * at the end are dropped, blank lines between items are kept as empty items. Text without
 * any item means "no list" and yields null.
 */
export function textToList(text: string): (string | string[])[] | null {
	const lines = text.split(/\r\n|\r|\n/);
	while (lines.length && lines[lines.length - 1] === "") lines.pop();
	return lines.length ? lines.map((line) => (line.includes("\t") ? line.split("\t") : line)) : null;
}

/**
 * ColumnWidths (`"72 pt;;1 in"`) → the width of each entry in points; undefined for a column
 * whose width the control works out (an empty entry or -1). Points unless `in` or `cm` follows.
 */
export function parseColumnWidths(text: unknown): (number | undefined)[] {
	if (typeof text !== "string" || !text.trim()) return [];
	return text.split(";").map((entry) => {
		const m = /^\s*(-?[\d.]+)\s*(pt|in|cm)?\s*$/i.exec(entry);
		if (!m) return undefined;
		const value = Number(m[1]);
		if (!Number.isFinite(value) || value < 0) return undefined;
		const unit = (m[2] ?? "pt").toLowerCase();
		return unit === "in" ? value * 72 : unit === "cm" ? (value * 72) / 2.54 : value;
	});
}

/**
 * Widths in points of the columns a list shows: ColumnCount of them (-1: as many as the
 * widest item has), each as ColumnWidths gives it; the others share what is left of
 * `available`, 72pt each at least (as MSForms does).
 */
export function columnLayout(columnCount: unknown, columnWidths: unknown, available: number, dataColumns: number): number[] {
	const count = typeof columnCount === "number" ? (columnCount < 0 ? Math.max(1, dataColumns) : columnCount) : 1;
	const given = parseColumnWidths(columnWidths);
	if (count <= 1 && given[0] === undefined) return count === 1 ? [available] : [];
	const widths = Array.from({ length: count }, (_, i) => given[i]);
	const fixed = widths.reduce<number>((sum, w) => sum + (w ?? 0), 0);
	const open = widths.filter((w) => w === undefined).length;
	const share = open ? Math.max(72, (available - fixed) / open) : 0;
	return widths.map((w) => w ?? share);
}

/**
 * The column whose text a ComboBox shows (and a ListBox's Text holds): TextColumn n is the
 * n-th; -1 (the default) the first one that is wider than 0. -1 when it is the row number.
 */
export function textColumnIndex(textColumn: unknown, widths: number[]): number {
	const n = typeof textColumn === "number" ? textColumn : -1;
	if (n === 0) return -1;
	if (n > 0) return n - 1;
	const first = widths.findIndex((w) => w > 0);
	return first < 0 ? 0 : first;
}
